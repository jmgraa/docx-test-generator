import re
from pathlib import Path
from string import ascii_uppercase

from docx import Document
from docx.shared import Pt

FONT_NAME = "Calibri"
NORMAL_SIZE = Pt(12)
SMALL_SIZE = Pt(8)
SCRIPT_PATTERN = re.compile(r"(<sup>.*?</sup>|<sub>.*?</sub>)", re.IGNORECASE)


def create_document(header_text, intro):
    doc = Document()

    header = doc.sections[0].header
    header_para = header.paragraphs[0]
    run_h = header_para.add_run(header_text)
    run_h.font.size = SMALL_SIZE
    run_h.font.name = FONT_NAME

    _add_paragraphs_to_document(doc, intro.split("\n"))

    return doc


def add_question_on_document(doc, index, question, answers):
    question_and_answers = [
        f"{index}) {question}",
        (" " * 5).join(
            [f"{ascii_uppercase[i]}) {a["text"]}" for i, a in enumerate(answers)]
        ),
    ]

    _add_paragraphs_to_document(doc, question_and_answers)
    doc.add_paragraph()

    return doc


def save_document(doc, index, dir):
    path = Path(dir) / f"{index}.docx"
    doc.save(path)


def save_dat(dat, index, dir):
    path = Path(dir) / f"k{index}.dat"
    with open(path, "w") as f:
        f.write("\n".join(dat))


def _add_paragraphs_to_document(doc, paragraphs):
    for text in paragraphs:
        p = doc.add_paragraph()
        _add_formatted_text(p, text)


def _add_formatted_text(paragraph, text):
    parts = SCRIPT_PATTERN.split(text)

    for part in parts:
        if not part:
            continue

        if part.lower().startswith("<sup>") and part.lower().endswith("</sup>"):
            content = part[5:-6]
            run = paragraph.add_run(content)
            run.font.superscript = True
        elif part.lower().startswith("<sub>") and part.lower().endswith("</sub>"):
            content = part[5:-6]
            run = paragraph.add_run(content)
            run.font.subscript = True
        else:
            run = paragraph.add_run(part)

        run.font.size = NORMAL_SIZE
        run.font.name = FONT_NAME
