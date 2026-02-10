import random
from pathlib import Path

from builder import add_question_on_document, create_document, save_dat, save_document
from parsing import parse_exam_file
from shuffling import shuffle_answers


def mix_tests(input_params):
    questions = parse_exam_file(input_params.file_path, input_params.output_dir)
    number_of_questions = len(questions)

    docx_dir = Path(input_params.output_dir) / "Testy"
    dat_dir = Path(input_params.output_dir) / "KluczeDAT"
    docx_dir.mkdir()
    dat_dir.mkdir()

    for i in range(1, input_params.copy_number + 1):
        yield i
        shuffled_questions = random.sample(questions, number_of_questions)
        datfile = []
        docx = create_document(
            f"{input_params.title}, {input_params.date}{" " * 5}Test {str(i).zfill(3)}",
            input_params.intro,
        )

        for q_num, q in enumerate(shuffled_questions):
            shuffled_answers, order = shuffle_answers(q)
            docx = add_question_on_document(docx, q_num + 1, q, shuffled_answers)
            datfile.append(f"{q.id} {" ".join(str(x + 1) for x in order)}")

        save_document(docx, i, docx_dir)
        save_dat(datfile, i, dat_dir)
