from pathlib import Path
from string import ascii_uppercase

from docx import Document
from docx.shared import Pt

FONT_NAME = "Calibri"
NORMAL_SIZE = Pt(12)
SMALL_SIZE = Pt(8)
PARAGRAPHS = [
    "PRZED PRZYSTĄPIENIEM DO WYPEŁNIANIA FORMULARZA ODPOWIEDZI PROSZĘ PRZECZYTAĆ INSTRUKCJĘ. PROSZĘ PAMIĘTAĆ O CZYTELNYM PODPISANIU FORMULARZA ODPOWIEDZI W PRAWYM-DOLNYM ROGU.",
    "PROSZĘ PAMIĘTAĆ O POPRAWNYM OZNACZENIU NUMERU TESTU I NUMERU IDENTYFIKACYJNEGO STUDENTA",
]


def create_document(header_text):
    doc = Document()

    header = doc.sections[0].header
    header_para = header.paragraphs[0]
    run_h = header_para.add_run(header_text)
    run_h.font.size = SMALL_SIZE
    run_h.font.name = FONT_NAME

    _add_paragraphs_to_document(doc, PARAGRAPHS)

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


def save_document(doc, index):
    path = Path("docs") / f"{index}.docx"
    doc.save(path)


def save_dat(dat, index):
    path = Path("keys") / f"k{index}.dat"
    with open(path, "w") as f:
        f.write("\n".join(dat))


def _add_paragraphs_to_document(doc, paragraphs):
    for text in paragraphs:
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.font.size = NORMAL_SIZE
        run.font.name = FONT_NAME
