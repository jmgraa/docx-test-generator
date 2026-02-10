import re
from pathlib import Path
from string import ascii_uppercase

from docx import Document
from docx.shared import Emu, Pt

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
    p_q = doc.add_paragraph()
    _add_formatted_text(p_q, f"{index}) {question.content}")
    _add_picture(question.images, doc)

    for i, answer in enumerate(answers):
        label = ascii_uppercase[i]

        p_a = doc.add_paragraph()
        _add_formatted_text(p_a, f"{label}) {answer['text']}")
        _add_picture(answer["images"], doc)

    doc.add_paragraph()

    return doc


def _add_picture(images, doc):
    for image_info in images:
        img_path = image_info.get("path")
        width_emu = image_info.get("width_emu")
        height_emu = image_info.get("height_emu")

        if not img_path:
            continue

        p_a_img = doc.add_paragraph()
        run_a_img = p_a_img.add_run()
        try:
            if width_emu and height_emu:
                run_a_img.add_picture(
                    str(img_path),
                    width=Emu(width_emu),
                    height=Emu(height_emu),
                )
            else:
                run_a_img.add_picture(str(img_path))
        except Exception:
            p_a_img.add_run(f"[OBRAZEK: {Path(img_path).name}]")


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
