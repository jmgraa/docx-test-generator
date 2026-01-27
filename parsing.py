from string import ascii_uppercase
from docx import Document

class ExamQuestion:
    def __init__(self, q_id, content):
        self.id = q_id
        self.content = content
        self.answers = []

    def add_answer(self, label, is_correct, text):
        self.answers.append(
            {"label": label, "is_correct": is_correct == "1", "text": text}
        )

    def __repr__(self):
        return f"Q{self.id}: {self.content[:40]}... ({len(self.answers)} ans)"


def get_formatted_text(cell):
    full_text = []
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            text = run.text
            if not text:
                continue

            if run.font.superscript:
                full_text.append(f"^{{{text}}}")
            elif run.font.subscript:
                full_text.append(f"_{{{text}}}")
            else:
                full_text.append(text)

        full_text.append(" ")

    return "".join(full_text).strip()


def parse_exam_file(file_path):
    doc = Document(file_path)
    questions = []
    current_q = None

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells

            if len(cells) != 5:
                continue

            col_text = [cells[i].text.strip() for i in range(3)]
            content_text = get_formatted_text(cells[4])

            if col_text[0].isdigit():
                current_q = ExamQuestion(col_text[0], content_text)
                questions.append(current_q)
            elif current_q and col_text[1] in ascii_uppercase:
                current_q.add_answer(col_text[1], col_text[2], content_text)

    return questions


file_path = "egzamin.docx"
questions = parse_exam_file(file_path)
for q in questions:
    print(q)
    for ans in q.answers:
        print(f"   {ans['label']}: {ans['text']}")
