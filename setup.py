import random
from pathlib import Path

from builder import add_question_on_document, create_document, save_dat, save_document
from parsing import parse_exam_file
from shuffling import shuffle_answers

file_path = "egzamin3.docx"
title = "Biofizyka Medyczna, LEK, Egzamin T01"
date = "2025-02-06"
copy_number = 3
questions = parse_exam_file(file_path)
number_of_questions = len(questions)

for dir in ["docs", "keys"]:
    Path(dir).mkdir()

for i in range(1, copy_number + 1):
    shuffled_questions = random.sample(questions, number_of_questions)
    datfile = []
    docx = create_document(f"{title}, {date}{" " * 5}Test {str(i).zfill(3)}")

    for q_num, q in enumerate(shuffled_questions):
        shuffled_answers, order = shuffle_answers(q)
        docx = add_question_on_document(docx, q_num + 1, q.content, shuffled_answers)
        datfile.append(f"{q.id} {" ".join(str(x + 1) for x in order)}")

    save_document(docx, i)
    save_dat(datfile, i)
